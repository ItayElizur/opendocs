using System;
using System.IO;

namespace OfficeAi.Shared
{
    // Post-hoc diagnostic tool (2026-08-24): a plain-text, append-only log for
    // manually reproducing bugs that have resisted blind fixing (the chart
    // RPC failure, in particular, after two rounds of unverified guesses).
    // Writes to a fixed, easy-to-find path so a user reproducing a bug can
    // locate and share the file without hunting for it - %TEMP% is a well
    // known Windows path nameable directly in Explorer's address bar.
    // Deliberately NOT wired into every tool call, only where a specific bug
    // needs step-by-step COM call tracing - this is a temporary debugging
    // aid, not a permanent logging subsystem.
    public static class DebugLog
    {
        public static readonly string LogPath = Path.Combine(Path.GetTempPath(), "AirchatOfficeDebug.log");
        private static readonly object Lock = new object();

        public static void Write(string message)
        {
            try
            {
                lock (Lock)
                {
                    File.AppendAllText(LogPath, DateTime.Now.ToString("HH:mm:ss.fff") + " " + message + Environment.NewLine);
                }
            }
            catch { /* logging must never break the tool it's diagnosing */ }
        }

        public static void WriteException(string context, Exception ex)
        {
            Write(context + " FAILED: " + ex.GetType().FullName + " HResult=0x" + ex.HResult.ToString("X8") + " Message=" + ex.Message);
            if (ex.StackTrace != null) Write("  StackTrace: " + ex.StackTrace.Replace("\r", "").Replace("\n", " | "));
            if (ex.InnerException != null) WriteException(context + " (inner)", ex.InnerException);
        }
    }
}
