using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    // Tool-level orchestration for everything in this add-in that needs
    // Exchange/EWS reachable, separate from the pure-COM tools elsewhere.
    // Low-level EWS wire calls live in OutlookEws.cs; this file resolves
    // which account/endpoint to use and turns a raw EWS result into a
    // ToolResult. On-prem Exchange only - see OutlookTools.Ews.cs.md for
    // the error-vs-fallback posture this file follows throughout.
    public static partial class OutlookTools
    {
        // Resolves a name/email fragment through EWS ResolveName (server-side
        // ANR over Contacts then the GAL), off the UI thread - the object
        // model can't do a multi-result directory search. See
        // OutlookTools.Ews.cs.md for why (includes the prior implementation's
        // failure mode).
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
            IReadOnlyList<ContactMatch> matches;
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

        // Cached once per process (mirrors OutlookEws.CachedUrl's lifetime).
        // _workWeekResolved latches true only on a SUCCESSFUL lookup - see
        // OutlookTools.Ews.cs.md for why a transient failure must not
        // permanently disable it for the rest of the Outlook session.
        private static OutlookEws.WorkWeekInfo? _cachedWorkWeek;
        private static bool _workWeekResolved;

        internal static readonly HashSet<DayOfWeek> FallbackWorkDays = new HashSet<DayOfWeek>
        {
            DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday,
        };

        // Best-effort: the real work week/hours come from EWS's
        // GetUserAvailability (no COM equivalent exists). Any failure
        // returns null so find_meeting_slots falls back to a COM-only
        // default instead of failing - unlike search_contacts, EWS here is
        // an enhancement, not the only way to do the job. See
        // OutlookTools.Ews.cs.md.
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

        // Shared by every EWS-dependent tool: resolves the endpoint once
        // (OutlookEws.CachedUrl short-circuits every later call, process-
        // wide). Always throws InvalidOperationException on failure - see
        // OutlookTools.Ews.cs.md for how the two callers use that
        // differently.
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
