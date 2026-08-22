using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CustomToolsForDoganium.Native;

/// <summary>
/// WH_KEYBOARD_LL ile sistem genelinde tuş basımlarını dinler ve her fiziksel basış için
/// (OS auto-repeat'i eleyerek) tek bir <see cref="KeyDown"/> olayı üretir.
/// Callback, Windows tarafından senkron çağrılır; burada YALNIZCA hızlı, engellemeyen iş
/// yapılmalıdır (bkz. <see cref="HookCallback"/>) — asıl kısayol mantığı bu sınıfın dışında,
/// UI thread'ine devredilerek çalıştırılır.
/// </summary>
internal sealed class KeyboardHook : IDisposable
{
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_KEYUP = 0x0101;
    private const int WM_SYSKEYUP = 0x0105;

    // Delege'yi bir alanda canlı tutuyoruz; aksi halde native taraf hâlâ referans tutarken
    // GC tarafından toplanıp callback'in çökmesine yol açabilir.
    private readonly LowLevelKeyboardProc _proc;
    private readonly HashSet<Keys> _keysCurrentlyDown = new();
    private IntPtr _hookHandle = IntPtr.Zero;

    internal event Action<Keys> KeyDown;

    internal KeyboardHook()
    {
        _proc = HookCallback;
    }

    internal void Install()
    {
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule;

        _hookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _proc,
            GetModuleHandle(curModule?.ModuleName), 0);

        if (_hookHandle == IntPtr.Zero)
            throw new InvalidOperationException(
                "Klavye kancası (keyboard hook) kurulamadı. Win32 hata kodu: " + Marshal.GetLastWin32Error());
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            if (nCode >= 0)
            {
                var key = (Keys)Marshal.ReadInt32(lParam);
                var message = wParam.ToInt32();

                switch (message)
                {
                    case WM_KEYDOWN or WM_SYSKEYDOWN:
                    {
                        // HashSet.Add, tuş zaten basılıyken (OS auto-repeat) false döner;
                        // böylece tek fiziksel basışta tek olay garanti edilir (eski
                        // "_wasKeyDown" edge-trigger mantığının event-driven karşılığı).
                        if (_keysCurrentlyDown.Add(key))
                            KeyDown?.Invoke(key);
                        break;
                    }
                    case WM_KEYUP or WM_SYSKEYUP:
                        _keysCurrentlyDown.Remove(key);
                        break;
                }
            }
        }
        catch (Exception ex)
        {
            // KRİTİK: Bu callback Windows tarafından native olarak çağrılıyor. Buradan
            // dışarı sızan HERHANGİ bir .NET exception'ı, Application.ThreadException'a
            // uğramadan işlemi anında çökertir. Bu yüzden burada asla exception
            // fırlatılmasına izin verilmez — sadece loglanır ve callback normal döner.
            Console.WriteLine("[KeyboardHook] Hata: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }

        // Diğer uygulamaların/kancaların girdiyi almasını asla engellemiyoruz.
        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    public void Dispose()
    {
        if (_hookHandle == IntPtr.Zero) return;
        UnhookWindowsHookEx(_hookHandle);
        _hookHandle = IntPtr.Zero;
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
}