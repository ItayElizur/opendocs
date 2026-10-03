## EwsAutodiscoverXml

```
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
```
