using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace OutlookAiAddIn
{
    // The non-COM calls in the Outlook add-in - the raw EWS wire layer.
    // Tool-facing orchestration built on top of this (which account/endpoint
    // to use, turning a result into a ToolResult) lives in
    // OutlookTools.Ews.cs, not here; this file only wraps the EWS Managed API
    // itself. Two operations: ResolveNamesAsync (search_contacts - a
    // server-side Ambiguous Name Resolution over Contacts then the GAL,
    // exactly what the native Address Book dialog does and what the
    // mcp-outlook reference's account.protocol.resolve_names does) and
    // GetWorkingHoursAsync (find_meeting_slots' work-week default - see its
    // own comment). Both are deliberately NOT Microsoft.Office.Interop.Outlook:
    // the COM object model can do neither a multi-result directory search nor
    // expose a mailbox's configured work week, and EWS is plain HTTP with no
    // STA affinity so both run off the UI thread (Task.Run) and never freeze
    // Outlook.
    //
    // Auth is ExchangeService.UseDefaultCredentials (Windows Integrated Auth as
    // the signed-in user) - the .NET equivalent of mcp-outlook's auth_type=sspi.
    // No stored credentials, no impersonation. On-prem Exchange only.
    //
    // This file has no `using Microsoft.Office.Interop.Outlook`: the two
    // namespaces collide on Contact, EmailAddress, Folder, Task, Item, ...
    internal static class OutlookEws
    {
        private const int TimeoutMs = 15000;

        static OutlookEws()
        {
            // EWS over HTTPS to on-prem Exchange fails with "The underlying connection
            // was closed: An unexpected error occurred on a send" when the process's
            // default security protocol doesn't offer TLS 1.2 - the .NET Framework
            // default and the EWS Managed API 2.2 (2014) vintage can both leave it off.
            // OR it in without disturbing anything already enabled.
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch { }
        }

        // Resolved once per process. Written on the UI thread after the first
        // successful discovery (the agent loop runs tool calls sequentially, so
        // there is never a concurrent writer).
        internal static Uri CachedUrl { get; set; }

        private static bool AllowHttpsRedirection(string redirectionUrl)
        {
            return redirectionUrl != null &&
                   redirectionUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
        }

        private static Ews.ExchangeService NewService(Uri url)
        {
            var svc = new Ews.ExchangeService(Ews.ExchangeVersion.Exchange2010_SP2)
            {
                UseDefaultCredentials = true, // do NOT also set Credentials
                Timeout = TimeoutMs,
                Url = url,
            };
            return svc;
        }

        // Autodiscover fallback for when Account.AutoDiscoverXml was empty or
        // unparseable. Runs the AD SCP lookup + HTTP off the UI thread.
        public static Task<Uri> DiscoverUrlAsync(string smtp)
        {
            return Task.Run(() =>
            {
                var svc = new Ews.ExchangeService(Ews.ExchangeVersion.Exchange2010_SP2)
                {
                    UseDefaultCredentials = true,
                    Timeout = TimeoutMs,
                };
                svc.AutodiscoverUrl(smtp, AllowHttpsRedirection);
                return svc.Url;
            });
        }

        public static Task<IReadOnlyList<KeyValuePair<string, string>>> ResolveNamesAsync(Uri url, string query)
        {
            return Task.Run(() => ResolveNames(url, query));
        }

        internal struct WorkWeekInfo
        {
            public HashSet<DayOfWeek> Days;
            public int StartHour;
            public int EndHour;
        }

        // GetUserAvailability's WorkingHours is EWS's documented, server-side
        // source for a mailbox's configured work days/hours - the same data
        // Outlook itself uses to shade "outside working hours" in the
        // scheduling assistant. Unlike Outlook's local Calendar Options
        // dialog, this has no COM equivalent at all: confirmed via .NET
        // reflection against the referenced Microsoft.Office.Interop.Outlook
        // PIA that no Application.CalendarOptions property (or any
        // WorkDay*/FirstDayOfWeek member) exists anywhere in that assembly,
        // so this EWS call is the only real, non-hardcoded source for it.
        public static Task<WorkWeekInfo?> GetWorkingHoursAsync(Uri url, string smtp)
        {
            return Task.Run(() => GetWorkingHours(url, smtp));
        }

        private static WorkWeekInfo? GetWorkingHours(Uri url, string smtp)
        {
            Ews.ExchangeService svc = NewService(url);
            var window = new Ews.TimeWindow(DateTime.Today, DateTime.Today.AddDays(7));
            Ews.GetUserAvailabilityResults results;
            try
            {
                results = svc.GetUserAvailability(
                    new[] { new Ews.AttendeeInfo(smtp) },
                    window,
                    Ews.AvailabilityData.FreeBusy);
            }
            catch (Ews.ServiceResponseException)
            {
                return null;
            }

            foreach (Ews.AttendeeAvailability a in results.AttendeesAvailability)
            {
                if (a.WorkingHours == null) continue;
                var days = new HashSet<DayOfWeek>();
                foreach (Ews.DayOfTheWeek d in a.WorkingHours.DaysOfTheWeek)
                {
                    switch (d)
                    {
                        case Ews.DayOfTheWeek.Sunday: days.Add(DayOfWeek.Sunday); break;
                        case Ews.DayOfTheWeek.Monday: days.Add(DayOfWeek.Monday); break;
                        case Ews.DayOfTheWeek.Tuesday: days.Add(DayOfWeek.Tuesday); break;
                        case Ews.DayOfTheWeek.Wednesday: days.Add(DayOfWeek.Wednesday); break;
                        case Ews.DayOfTheWeek.Thursday: days.Add(DayOfWeek.Thursday); break;
                        case Ews.DayOfTheWeek.Friday: days.Add(DayOfWeek.Friday); break;
                        case Ews.DayOfTheWeek.Saturday: days.Add(DayOfWeek.Saturday); break;
                        // Day/Weekday/WeekendDay are input-only aggregate
                        // values per the enum's own shape (confirmed via
                        // reflection) - never expected back from the server,
                        // so deliberately not expanded here.
                    }
                }
                if (days.Count == 0) continue;
                return new WorkWeekInfo
                {
                    Days = days,
                    StartHour = (int)a.WorkingHours.StartTime.TotalHours,
                    EndHour = (int)a.WorkingHours.EndTime.TotalHours,
                };
            }
            return null;
        }

        private static IReadOnlyList<KeyValuePair<string, string>> ResolveNames(Uri url, string query)
        {
            var results = new List<KeyValuePair<string, string>>();

            Ews.ExchangeService svc = NewService(url);
            Ews.NameResolutionCollection col;
            try
            {
                col = svc.ResolveName(query, Ews.ResolveNameSearchLocation.ContactsThenDirectory, true);
            }
            catch (Ews.ServiceResponseException)
            {
                // ErrorNameResolutionNoResults / NoMailbox surface here on some
                // servers rather than as an empty collection - treat as "no matches".
                return results;
            }

            foreach (Ews.NameResolution nr in col)
            {
                string email = SmtpFrom(nr);
                if (string.IsNullOrEmpty(email)) continue; // mirror mcp-outlook: drop entries with no address

                string name = DisplayNameFrom(nr, email);
                results.Add(new KeyValuePair<string, string>(name, email));
            }

            return results;
        }

        private static string SmtpFrom(Ews.NameResolution nr)
        {
            Ews.EmailAddress mb = nr.Mailbox;
            if (mb != null && !string.IsNullOrEmpty(mb.Address) && mb.Address.Contains("@") &&
                (string.IsNullOrEmpty(mb.RoutingType) || mb.RoutingType.Equals("SMTP", StringComparison.OrdinalIgnoreCase)))
            {
                return mb.Address;
            }

            // GAL hits often carry an X500/legacyExchangeDN Address (RoutingType
            // "EX", no "@"); fall back to the resolved contact's own addresses.
            Ews.Contact contact = nr.Contact;
            if (contact != null && contact.EmailAddresses != null)
            {
                foreach (Ews.EmailAddressKey key in new[]
                {
                    Ews.EmailAddressKey.EmailAddress1,
                    Ews.EmailAddressKey.EmailAddress2,
                    Ews.EmailAddressKey.EmailAddress3,
                })
                {
                    Ews.EmailAddress ea;
                    if (contact.EmailAddresses.TryGetValue(key, out ea) &&
                        ea != null && !string.IsNullOrEmpty(ea.Address) && ea.Address.Contains("@"))
                    {
                        return ea.Address;
                    }
                }
            }

            return null;
        }

        private static string DisplayNameFrom(Ews.NameResolution nr, string emailFallback)
        {
            Ews.Contact contact = nr.Contact;
            if (contact != null)
            {
                if (!string.IsNullOrEmpty(contact.DisplayName)) return contact.DisplayName;
                string given = contact.GivenName ?? "";
                string surname = contact.Surname ?? "";
                string joined = (given + " " + surname).Trim();
                if (joined.Length > 0) return joined;
            }
            if (nr.Mailbox != null && !string.IsNullOrEmpty(nr.Mailbox.Name)) return nr.Mailbox.Name;
            return emailFallback;
        }

        // Timeouts surface inconsistently across EWS Managed API paths - as a
        // bare TimeoutException, or a WebException(Timeout), or either wrapped in
        // a ServiceRequestException. Walk the whole inner chain.
        public static bool IsTimeout(Exception ex)
        {
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                if (e is TimeoutException) return true;
                var web = e as WebException;
                if (web != null && web.Status == WebExceptionStatus.Timeout) return true;
            }
            return false;
        }
    }
}
