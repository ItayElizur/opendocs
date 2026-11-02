using System;
using System.Runtime.InteropServices;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Repairs a mouse pointer that is already invisible on the calling UI thread. Seen in Excel: when a tool that opens
    /// a dialog is sent with the mouse over the chat pane, the thread's cursor is null (sometimes also with a negative
    /// ShowCursor count) by the time the dialog opens, and stays hidden until the pointer re-enters the pane. Call just
    /// before showing the dialog. See HiddenPointer.cs.md.
    /// </summary>
    public static class HiddenPointer
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct CursorInfo { public int cbSize; public int flags; public IntPtr hCursor; public int x; public int y; }
        [DllImport("user32.dll")] private static extern bool GetCursorInfo(ref CursorInfo ci);
        [DllImport("user32.dll")] private static extern int ShowCursor(bool show);
        [DllImport("user32.dll")] private static extern IntPtr SetCursor(IntPtr cursor);
        [DllImport("user32.dll")] private static extern IntPtr LoadCursor(IntPtr instance, int name);

        private const int IdcArrow = 32512;

        private static int ShowCursorCount()
        {
            ShowCursor(true); // the API only reports the counter when changing it, so bump and undo
            return ShowCursor(false);
        }

        /// <summary>Winds a negative ShowCursor count back to 0 and replaces a null cursor with the arrow. Never throws.</summary>
        public static void Reset()
        {
            try
            {
                var ci = new CursorInfo { cbSize = Marshal.SizeOf(typeof(CursorInfo)) };
                bool ok = GetCursorInfo(ref ci);
                int count = ShowCursorCount();
                int steps = 0;
                while (count < 0 && steps < 20) { ShowCursor(true); count = ShowCursorCount(); steps++; }
                bool nullCursor = ok && ci.hCursor == IntPtr.Zero;
                if (nullCursor) SetCursor(LoadCursor(IntPtr.Zero, IdcArrow));
                if (steps > 0 || nullCursor)
                    DebugLog.Write("reset hidden pointer (showCursor steps=" + steps + ", null cursor replaced=" + nullCursor + ")");
            }
            catch (Exception ex)
            {
                DebugLog.Write("hidden pointer reset failed: " + ex.Message);
            }
        }
    }
}
