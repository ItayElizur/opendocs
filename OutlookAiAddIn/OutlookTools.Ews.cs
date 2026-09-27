using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    // Every tool-level thing in this add-in that needs Exchange/EWS to be
    // reachable, in one place, separate from the pure-COM tools in the other
    // OutlookTools.*.cs files (which only need the already-running, already-
    // signed-in Outlook client and work the same regardless of mailbox type
    // or network state). Low-level EWS wire calls themselves live one layer
    // further down, in OutlookEws.cs; this file is the tool-facing
    // orchestration on top of it - resolving which account/endpoint to use,
    // and turning a raw EWS result into a ToolResult or a plain value another
    // tool can consume.
    //
    // On-prem Exchange only, same posture everywhere in this file: no
    // Exchange account in the profile, or EWS unreachable, is either a clear
    // IsError result (search_contacts, where EWS is the only way to do the
    // job) or a graceful fallback to a COM-only default (find_meeting_slots'
    // work-week lookup, where a COM-derived value already exists to fall
    // back to). Never a silent partial success.
    public static partial class OutlookTools
    {
        // search_contacts resolves a name/email fragment through EWS ResolveName
        // (server-side ANR over Contacts then the GAL) - the same thing the
        // native Address Book does, and what the mcp-outlook reference does. It
        // is one of two tools that aren't Outlook COM: the object model can't do
        // a multi-result directory search, and running EWS off the UI thread
        // (OutlookEws.ResolveNamesAsync -> Task.Run) is what keeps Outlook
        // responsive during the call. The previous implementation walked every
        // contact folder in every store on the UI thread and froze/crashed
        // Outlook. See docs/ai-tool-surface.md and OutlookEws.cs.
        private static async Task<ToolResult> SearchContactsAsync(JsonElement input)
        {
            string query = ReqStr(input, "query");
            int limit = Math.Max(1, Int(input, "limit", 10));

            Uri url;
            try
            {
                url = await ResolveEwsUrlAsync();
            }
            catch (InvalidOperationException ex)
            {
                return Err(ex.Message, "search_contacts");
            }

            // ResolveName, off the UI thread.
            IReadOnlyList<KeyValuePair<string, string>> matches;
            try
            {
                matches = await OutlookEws.ResolveNamesAsync(url, query);
            }
            catch (Exception ex) when (OutlookEws.IsTimeout(ex))
            {
                DebugLog.WriteException("search_contacts ResolveName timeout", ex);
                return Err("The Exchange contact search timed out after 15s. Try again, or check your network / VPN connection.", "search_contacts");
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException ex)
            {
                DebugLog.WriteException("search_contacts ResolveName", ex);
                return Err("Exchange rejected the contact search: " + ex.Message +
                           " (Windows authentication to Exchange may have failed - are you on the domain network?)", "search_contacts");
            }
            catch (WebException ex)
            {
                DebugLog.WriteException("search_contacts ResolveName WebException", ex);
                return Err(ex.Status == WebExceptionStatus.ProtocolError
                    ? "Windows authentication to Exchange failed (are you connected to the domain network / VPN?)."
                    : "Could not reach the Exchange server (" + ex.Status + ").", "search_contacts");
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("search_contacts ResolveName (unexpected)", ex);
                return Err("Contact search failed: " + ex.Message, "search_contacts");
            }

            // Format (pure, back on the UI thread).
            return new ToolResult
            {
                Output = ContactSearchFormat.Format(matches, query, limit),
                Summary = "search_contacts",
            };
        }

        // Cached once per process, same lifetime/posture as OutlookEws.CachedUrl
        // - the mailbox's configured work days/hours don't change mid-session,
        // and each lookup is a network round trip via EWS. Mirrors CachedUrl's
        // posture exactly: _workWeekResolved only latches true on a SUCCESSFUL
        // lookup, never on failure - a transient EWS/network blip on the first
        // find_meeting_slots call must not permanently disable the real
        // work-week lookup (falling back to Sun-Thu/9-18) for the rest of the
        // Outlook session, which can run for days.
        private static OutlookEws.WorkWeekInfo? _cachedWorkWeek;
        private static bool _workWeekResolved;

        internal static readonly HashSet<DayOfWeek> FallbackWorkDays = new HashSet<DayOfWeek>
        {
            DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday,
        };

        // Best-effort: the real work week/hours come from EWS's
        // GetUserAvailability (see OutlookEws.GetWorkingHoursAsync's own
        // comment for why - no COM equivalent exists). On-prem Exchange only,
        // same as search_contacts; any failure (no Exchange account, EWS
        // unreachable, etc.) returns null so the caller (find_meeting_slots)
        // can fall back to a COM-only default, rather than failing the tool -
        // unlike search_contacts, where EWS isn't optional, here it's an
        // enhancement over an already-working default.
        internal static async Task<OutlookEws.WorkWeekInfo?> ResolveWorkWeekAsync()
        {
            if (_workWeekResolved) return _cachedWorkWeek;
            try
            {
                Uri url = await ResolveEwsUrlAsync(); // throws on any failure - caught below, not propagated
                string smtp = FindExchangeAccountInfo().smtp;
                OutlookEws.WorkWeekInfo? result = await OutlookEws.GetWorkingHoursAsync(url, smtp);
                if (result.HasValue)
                {
                    // Only a successful lookup is cached/latched - see the
                    // fields' own comment above.
                    _cachedWorkWeek = result;
                    _workWeekResolved = true;
                }
                return result;
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("ResolveWorkWeekAsync", ex);
                return null;
            }
        }

        // Shared by every EWS-dependent tool: resolve the endpoint once
        // (OutlookEws.CachedUrl short-circuits every call after the first,
        // process-wide, regardless of which tool triggered the discovery).
        // Always throws InvalidOperationException on failure, with a message
        // specific enough for search_contacts to surface directly -
        // ResolveWorkWeekAsync above just catches and discards it (a graceful
        // null is all it wants).
        private static async Task<Uri> ResolveEwsUrlAsync()
        {
            Uri url = OutlookEws.CachedUrl;
            if (url != null) return url;

            var info = FindExchangeAccountInfo(); // throws InvalidOperationException itself if no Exchange account
            string parsed = EwsAutodiscoverXml.ParseEwsUrl(info.autodiscoverXml);
            if (!string.IsNullOrEmpty(parsed))
            {
                try { url = new Uri(parsed); }
                catch (UriFormatException) { url = null; }
            }

            if (url == null)
            {
                if (string.IsNullOrEmpty(info.smtp))
                    throw new InvalidOperationException("Could not determine the Exchange EWS endpoint: no autodiscover data and no SMTP address to probe.");
                try
                {
                    url = await OutlookEws.DiscoverUrlAsync(info.smtp);
                }
                catch (Exception ex)
                {
                    DebugLog.WriteException("ResolveEwsUrlAsync AutodiscoverUrl", ex);
                    throw new InvalidOperationException("Could not reach Exchange autodiscover to find the EWS endpoint (" + ex.Message + ").", ex);
                }
            }

            OutlookEws.CachedUrl = url; // back on the UI thread after the await
            return url;
        }

        // COM, UI thread. Picks the Exchange account to discover the EWS endpoint
        // and SMTP from: the one matching the signed-in user when there is a
        // match, else the first Exchange account in the profile. Shared by every
        // EWS-dependent tool above.
        private static (string smtp, string autodiscoverXml) FindExchangeAccountInfo()
        {
            Outlook.NameSpace ns = Ns;

            string currentSmtp = null;
            try { currentSmtp = SmtpOf(ns.CurrentUser.AddressEntry); }
            catch (Exception ex) { DebugLog.WriteException("FindExchangeAccountInfo CurrentUser SMTP", ex); }

            Outlook.Account chosen = null;
            Outlook.Accounts accounts = ns.Accounts;
            for (int i = 1; i <= accounts.Count; i++)
            {
                Outlook.Account a = accounts[i];
                if (a.AccountType != Outlook.OlAccountType.olExchange) continue;
                if (chosen == null) chosen = a;
                if (!string.IsNullOrEmpty(currentSmtp) &&
                    string.Equals(a.SmtpAddress, currentSmtp, StringComparison.OrdinalIgnoreCase))
                {
                    chosen = a;
                    break;
                }
            }

            if (chosen == null)
                throw new InvalidOperationException(
                    "This needs an on-prem Exchange mailbox in this Outlook profile. " +
                    "No Exchange account was found (Gmail / IMAP / POP profiles are not supported).");

            string smtp = !string.IsNullOrEmpty(chosen.SmtpAddress) ? chosen.SmtpAddress : currentSmtp;

            string xml = null;
            try { xml = chosen.AutoDiscoverXml; }
            catch (Exception ex) { DebugLog.WriteException("FindExchangeAccountInfo Account.AutoDiscoverXml", ex); }

            return (smtp, xml);
        }

        private static ToolResult Err(string message, string summary)
        {
            return new ToolResult { Output = message, IsError = true, Summary = summary };
        }
    }
}
