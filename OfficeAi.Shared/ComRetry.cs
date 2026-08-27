using System;

namespace OfficeAi.Shared
{
    /// <summary>
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
    /// </summary>
    public static class ComRetry
    {
        public static readonly int[] TransientHResults =
        {
            unchecked((int)0x800706BE), // RPC_S_CALL_FAILED - "The remote procedure call failed."
            unchecked((int)0x8001010A), // RPC_E_SERVERCALL_RETRYLATER - "The message filter indicated that the application is busy."
            unchecked((int)0x800706BA), // RPC_S_SERVER_UNAVAILABLE
        };

        public static bool IsTransient(int hResult)
        {
            return Array.IndexOf(TransientHResults, hResult) >= 0;
        }

        // Every attempt (success, retried failure, AND a non-retried failure)
        // is logged - this is what surfaces the REAL exception detail from a
        // live repro instead of leaving it to guesswork.
        public static void Run(Action action, string label = "ComRetry")
        {
            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    DebugLog.Write(label + ": attempt " + attempt + " starting");
                    action();
                    DebugLog.Write(label + ": attempt " + attempt + " SUCCEEDED");
                    return;
                }
                catch (Exception ex) when (attempt < maxAttempts && IsTransient(ex.HResult))
                {
                    DebugLog.WriteException(label + ": attempt " + attempt + " (transient, retrying)", ex);
                    System.Threading.Thread.Sleep(200 * attempt);
                }
                catch (Exception ex)
                {
                    // Either the last attempt, or an HResult not in the
                    // transient list - logged before rethrowing so the real
                    // failure is captured even when no more retries happen.
                    DebugLog.WriteException(label + ": attempt " + attempt + " (NOT retried - rethrowing)", ex);
                    throw;
                }
            }
        }
    }
}
