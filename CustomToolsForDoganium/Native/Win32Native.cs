using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CustomToolsForDoganium.Native;

/// <summary>
/// Win32 API çağrılarını saran, önplandaki pencere/işlem bilgisi ve tuş durumu
/// erişimi sağlayan tek yardımcı sınıf. Tüm P/Invoke tanımları burada toplanır.
/// </summary>
internal static class Win32Native
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(Keys vKey);

    [DllImport("user32.dll")]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    internal static bool IsKeyDown(Keys key) => (GetAsyncKeyState(key) & 0x8000) != 0;

    internal static bool TryGetForegroundWindow(out IntPtr hwnd)
    {
        hwnd = GetForegroundWindow();
        return hwnd != IntPtr.Zero;
    }

    internal static string GetWindowTitle(IntPtr hwnd)
    {
        var length = GetWindowTextLength(hwnd);
        var sb = new StringBuilder(length + 1);
        GetWindowText(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }

    /// <summary>Önplandaki pencerenin process adını (küçük harf) döner. Bulunamazsa null.</summary>
    internal static string GetForegroundProcessName()
    {
        if (!TryGetForegroundWindow(out var hwnd)) return null;
        return TryGetProcessInfo(hwnd, out var processName, out _, out _) ? processName.ToLower() : null;
    }

    /// <summary>Verilen pencere tutamacının ait olduğu process bilgilerini güvenli şekilde döner.</summary>
    internal static bool TryGetProcessInfo(IntPtr hwnd, out string processName, out int processId,
        out string executablePath)
    {
        processName = null;
        processId = 0;
        executablePath = null;

        GetWindowThreadProcessId(hwnd, out var pid);

        try
        {
            using var proc = Process.GetProcessById((int)pid);
            processName = proc.ProcessName;
            processId = (int)pid;
            executablePath = TryGetMainModulePath(proc);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string TryGetMainModulePath(Process proc)
    {
        try
        {
            return proc.MainModule?.FileName;
        }
        catch
        {
            // Bazı sistem/yönetici işlemlerinde MainModule erişimi kısıtlı olabilir.
            return null;
        }
    }
}