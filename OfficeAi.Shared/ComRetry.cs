using System;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Retry wrapper for Office COM calls that fail *transiently* (for reasons
    /// unrelated to the request being wrong, succeeding moments later).
    /// Allowlist-only - see <see cref="TransientHResults"/> and ComRetry.cs.md
    /// for the full rationale and history.
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
