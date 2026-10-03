using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeAi.Shared
{
    // Pure, COM-free extraction of the EWS endpoint URL from an Outlook
    // Autodiscover response. Lives here (not in OutlookAiAddIn) so it can be
    // unit tested without the Outlook PIA. Matches elements by local name
    // only, ignoring the response's schema namespace. See EwsAutodiscoverXml.cs.md.
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
