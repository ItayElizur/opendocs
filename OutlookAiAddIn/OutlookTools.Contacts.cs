using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        // search_contacts resolves a name/email fragment through EWS ResolveName
        // (server-side ANR over Contacts then the GAL) - the same thing the
        // native Address Book does, and what the mcp-outlook reference does. It
        // is the ONE tool that isn't Outlook COM: the object model can't do a
        // multi-result directory search, and running EWS off the UI thread
        // (OutlookEws.ResolveNamesAsync -> Task.Run) is what keeps Outlook
        // responsive during the call. The previous implementation walked every
        // contact folder in every store on the UI thread and froze/crashed
        // Outlook. See docs/ai-tool-surface.md and OutlookEws.cs.
        private static async Task<ToolResult> SearchContactsAsync(JsonElement input)
        {
            string query = ReqStr(input, "query");
            int limit = Math.Max(1, Int(input, "limit", 10));

            // 1) Resolve the EWS endpoint (cached for the process). The
            //    AutoDiscoverXml read is COM and stays on the UI thread; it is a
            //    cached property, effectively instant.
            Uri url = OutlookEws.CachedUrl;
            if (url == null)
            {
                string smtp;
                string autodiscoverXml;
                try
                {
                    var info = FindExchangeAccountInfo();
                    smtp = info.smtp;
                    autodiscoverXml = info.autodiscoverXml;
                }
                catch (InvalidOperationException ex)
                {
                    return Err(ex.Message);
                }

                string parsed = EwsAutodiscoverXml.ParseEwsUrl(autodiscoverXml);
                if (!string.IsNullOrEmpty(parsed))
                {
                    try { url = new Uri(parsed); }
                    catch (UriFormatException) { url = null; }
                }

                if (url == null)
                {
                    if (string.IsNullOrEmpty(smtp))
                        return Err("Could not determine the Exchange EWS endpoint: no autodiscover data and no SMTP address to probe.");
                    try
                    {
                        url = await OutlookEws.DiscoverUrlAsync(smtp);
                    }
                    catch (Exception ex)
                    {
                        DebugLog.WriteException("search_contacts AutodiscoverUrl", ex);
                        return Err("Could not reach Exchange autodiscover to find the EWS endpoint (" + ex.Message + ").");
                    }
                }

                OutlookEws.CachedUrl = url; // back on the UI thread after the await
            }

            // 2) ResolveName, off the UI thread.
            IReadOnlyList<KeyValuePair<string, string>> matches;
            try
            {
                matches = await OutlookEws.ResolveNamesAsync(url, query);
            }
            catch (Exception ex) when (OutlookEws.IsTimeout(ex))
            {
                DebugLog.WriteException("search_contacts ResolveName timeout", ex);
                return Err("The Exchange contact search timed out after 15s. Try again, or check your network / VPN connection.");
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException ex)
            {
                DebugLog.WriteException("search_contacts ResolveName", ex);
                return Err("Exchange rejected the contact search: " + ex.Message +
                           " (Windows authentication to Exchange may have failed - are you on the domain network?)");
            }
            catch (WebException ex)
            {
                DebugLog.WriteException("search_contacts ResolveName WebException", ex);
                return Err(ex.Status == WebExceptionStatus.ProtocolError
                    ? "Windows authentication to Exchange failed (are you connected to the domain network / VPN?)."
                    : "Could not reach the Exchange server (" + ex.Status + ").");
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("search_contacts ResolveName (unexpected)", ex);
                return Err("Contact search failed: " + ex.Message);
            }

            // 3) Format (pure, back on the UI thread).
            return new ToolResult
            {
                Output = ContactSearchFormat.Format(matches, query, limit),
                Summary = "search_contacts",
            };
        }

        // COM, UI thread. Picks the Exchange account to discover the EWS endpoint
        // and SMTP from: the one matching the signed-in user when there is a
        // match, else the first Exchange account in the profile.
        private static (string smtp, string autodiscoverXml) FindExchangeAccountInfo()
        {
            Outlook.NameSpace ns = Ns;

            string currentSmtp = null;
            try { currentSmtp = SmtpOf(ns.CurrentUser.AddressEntry); }
            catch (Exception ex) { DebugLog.WriteException("search_contacts CurrentUser SMTP", ex); }

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
                    "search_contacts needs an on-prem Exchange mailbox in this Outlook profile. " +
                    "No Exchange account was found (Gmail / IMAP / POP profiles are not supported for contact search).");

            string smtp = !string.IsNullOrEmpty(chosen.SmtpAddress) ? chosen.SmtpAddress : currentSmtp;

            string xml = null;
            try { xml = chosen.AutoDiscoverXml; }
            catch (Exception ex) { DebugLog.WriteException("search_contacts Account.AutoDiscoverXml", ex); }

            return (smtp, xml);
        }

        private static ToolResult Err(string message)
        {
            return new ToolResult { Output = message, IsError = true, Summary = "search_contacts" };
        }
    }
}
