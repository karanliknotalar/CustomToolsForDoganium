using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace CustomToolsForDoganium.Manager
{
    internal static class Win32Native
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        internal static extern short GetAsyncKeyState(int vKey);

        internal const int VK_CONTROL = 0x11;
        internal const int VK_Q = 0x51; // 'Q' tuşu
        internal const int VK_E = 0x45; // 'E' tuşu
        internal const int VK_B = 0x42; // 'B' tuşu
        internal const int VK_N = 0x4E; // 'N' tuşu
        internal const int VK_T = 0x54; // 'T' tuşu
        internal const int VK_OEM_2 = 0xBF; // Türkçe klavyede 'Ö' tuşu
        // Yeni kısayollar için buraya VK_ sabitleri eklenir

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