using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using CustomToolsForDoganium.Capture;
using CustomToolsForDoganium.Native;
using CustomToolsForDoganium.Notifications;

namespace CustomToolsForDoganium.Hotkeys;

/// <summary>
/// <see cref="KeyboardHook"/>'tan gelen tuş olaylarını kayıtlı <see cref="IHotkeyAction"/>'larla
/// ve yakalama (Ctrl+Shift+Tuş) kısayollarıyla eşleştirir. Eşleşme anındaki tespit hook thread'inde
/// (hızlı, birkaç p/invoke) yapılır; gerçek eylem ise <paramref name="uiThread"/> üzerinden
/// UI thread'ine devredilir — low-level hook callback'i asla bloklanmaz.
/// </summary>
internal sealed class GlobalHotkeyDispatcher(
    IEnumerable<IHotkeyAction> actions,
    IEnumerable<(Keys Key, InsuranceQueryType QueryType)> captureBindings,
    DoganiumCaptureService captureService,
    NotifyIcon notifyIcon,
    ISynchronizeInvoke uiThread)
{
    internal void OnKeyDown(Keys key)
    {
        if (!Win32Native.IsKeyDown(Keys.ControlKey)) return;

        if (Win32Native.IsKeyDown(Keys.ShiftKey))
        {
            foreach (var binding in captureBindings)
            {
                if (binding.Key != key) continue;
                Dispatch(() => captureService.Capture(binding.QueryType));
                return;
            }

            return;
        }

        foreach (var action in actions)
        {
            if (action.VirtualKey != key) continue;
            Dispatch(() => TryExecute(action));
            return;
        }
    }

    private void TryExecute(IHotkeyAction action)
    {
        var processName = Win32Native.GetForegroundProcessName();
        if (processName == null) return;

        var activeApp = ResolveTargetApp(processName);
        if (activeApp == TargetApp.None) return;
        if ((action.SupportedApps & activeApp) == 0) return;

        var clipboardText = Clipboard.ContainsText() ? Clipboard.GetText() : null;
        action.Execute(clipboardText, activeApp, notifyIcon);
    }

    /// <summary>İşi UI thread'inin mesaj kuyruğuna kuyruklar ve orada güvenli şekilde
    /// try/catch ile çalıştırır (eski while-loop'un hata yakalama davranışının karşılığı).</summary>
    private void Dispatch(Action work)
    {
        uiThread.BeginInvoke(new Action(() => RunSafely(work)), null);
    }

    private void RunSafely(Action work)
    {
        try
        {
            work();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            Console.WriteLine("Stack Trace: " + ex.StackTrace);
            NotificationService.ShowWarning(notifyIcon, "TEKRAR DENE! KISAYOL İŞLENİRKEN BİR HATA OLUŞTU.");
        }
    }

    private static TargetApp ResolveTargetApp(string processName)
    {
        if (processName.Contains("notepad++")) return TargetApp.NotepadPlusPlus;
        if (processName.Contains("excel")) return TargetApp.Excel;
        return TargetApp.None;
    }
}