using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CustomToolsForDoganium.Manager
{
    internal static class Win32Native
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        internal static extern short GetAsyncKeyState(Keys vKey);


        /// <summary>Önplandaki pencerenin process adını (küçük harf) döner. Bulunamazsa null.</summary>
        internal static string GetForegroundProcessName()
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return null;

            GetWindowThreadProcessId(hwnd, out var processId);

            try
            {
                using var proc = Process.GetProcessById((int)processId);
                return proc.ProcessName.ToLower();
            }
            catch
            {
                return null;
            }
        }
    }
}