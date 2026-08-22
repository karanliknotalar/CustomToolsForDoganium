using System;
using System.Runtime.InteropServices;

namespace CustomToolsForDoganium.Native;

/// <summary>Uygulamanın konsol penceresini gizleme/geri getirme ile ilgili Win32 çağrıları.</summary>
internal static class ConsoleWindowNative
{
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    private const int SW_HIDE = 0;
    private const int SW_RESTORE = 9;

    internal static IntPtr GetHandle() => GetConsoleWindow();

    internal static bool IsMinimized(IntPtr hwnd) => IsIconic(hwnd);

    internal static void Hide(IntPtr hwnd) => ShowWindow(hwnd, SW_HIDE);

    internal static void Restore(IntPtr hwnd) => ShowWindow(hwnd, SW_RESTORE);
}