using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeAi.Shared
{
    // Pure, COM-free extraction of the EWS endpoint URL from an Outlook
    // Autodiscover response (Outlook.Account.AutoDiscoverXml). Lives here (not in
    // OutlookAiAddIn) so it can be unit tested without the Outlook PIA. The
    // add-in feeds the returned string into ExchangeService.Url for
    // search_contacts' EWS ResolveName call.
    //
    // The POX response nests <Response>/<Account>/<Protocol> elements in the
    // http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a
    // namespace, but we match purely on local names so a schema/namespace
    // variation can't silently break discovery - a null return just falls the
    // caller through to ExchangeService.AutodiscoverUrl instead.
    public static class EwsAutodiscoverXml
    {
        // Returns the EWS asmx URL, or null when the XML is empty, malformed,
        // an <Error> response, or carries no usable protocol URL.
        public static string ParseEwsUrl(string autodiscoverXml)
        {
            if (string.IsNullOrWhiteSpace(autodiscoverXml)) return null;

            try
            {
                XDocument doc = XDocument.Parse(autodiscoverXml);

                string expr = null;
                string other = null;

                foreach (XElement protocol in doc.Descendants())
                {
                    if (protocol.Name.LocalName != "Protocol") continue;

                    string type = ChildValue(protocol, "Type");
                    string url = ChildValue(protocol, "EwsUrl");
                    if (string.IsNullOrEmpty(url)) url = ChildValue(protocol, "ASUrl");
                    if (string.IsNullOrEmpty(url)) continue;

                    if (string.Equals(type, "EXCH", StringComparison.OrdinalIgnoreCase))
                        return url; // internal endpoint - preferred on a domain-joined box
                    if (expr == null && string.Equals(type, "EXPR", StringComparison.OrdinalIgnoreCase))
                        expr = url;
                    if (other == null)
                        other = url;
                }

                return expr ?? other;
            }
            catch
            {
                return null;
            }
        }

        private static string ChildValue(XElement parent, string localName)
        {
            foreach (XElement child in parent.Elements())
                if (child.Name.LocalName == localName)
                {
                    string v = child.Value;
                    return string.IsNullOrWhiteSpace(v) ? null : v.Trim();
                }
            return null;
        }
    }
}
