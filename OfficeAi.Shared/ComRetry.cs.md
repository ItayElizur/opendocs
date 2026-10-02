## ComRetry

```
/// Retry wrapper for Office COM calls that fail *transiently* - i.e. for
/// reasons unrelated to the request being wrong, which succeed on a
/// retry moments later.
///
/// The case this exists for: writing chart data drives Office's embedded
/// chart workbook, an out-of-process OLE server (a hidden Excel). Under
/// rapid successive calls that server sometimes simply declines, with one
/// of the HRESULTs in <see cref="TransientHResults"/>. None of them mean
/// "you passed a bad range"; they mean "Office was busy, ask again".
///
/// Deliberately an ALLOWLIST, not a blanket catch: only those specific
/// HRESULTs retry, so a genuine logic error (bad range, missing property)
/// still fails immediately on the first attempt rather than being masked
/// behind three retries and ~600ms of delay.
///
/// Previously duplicated in WordTools.cs and PowerPointTools.cs - the two
/// copies had drifted only in whether they took a `label`. The shared
/// version keeps the label (defaulted), since it is what makes the debug
/// log readable when several charts are written in one session.
```
