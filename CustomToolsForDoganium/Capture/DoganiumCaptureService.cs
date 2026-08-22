using System;
using System.Media;
using System.Windows.Forms;
using CustomToolsForDoganium.Native;
using CustomToolsForDoganium.Notifications;

namespace CustomToolsForDoganium.Capture;

/// <summary>Ctrl+Shift+C/X/Z tetiklendiğinde: önplandaki pencerenin gerçekten bir Doğanium
/// sorgulama ekranı olduğunu doğrular, UI Automation ile metnini okur, uygun formatta
/// özetler ve panoya yazar. Eski Program.cs'teki CaptureText_SafeWindowsOnly + ProcessAndFormatText
/// akışının tek sorumluluğa indirgenmiş hali.</summary>
internal sealed class DoganiumCaptureService(NotifyIcon notifyIcon)
{
    internal void Capture(InsuranceQueryType queryType)
    {
        Console.WriteLine();
        Console.WriteLine("=== YENİ YAKALAMA DENEMESİ ===");

        if (!Win32Native.TryGetForegroundWindow(out var hwnd))
        {
            NotificationService.ShowWarning(notifyIcon, "Aktif pencere bulunamadı");
            return;
        }

        Console.WriteLine($"HWND: {hwnd}");

        var windowTitle = Win32Native.GetWindowTitle(hwnd);
        if (!windowTitle.Contains("Sorgulama Ekranı"))
        {
            NotificationService.ShowWarning(notifyIcon, "Aktif pencere Dogaium sorgulama ekranı değil!");
            return;
        }

        Console.WriteLine($"Pencere Başlığı: '{windowTitle}'");

        if (!Win32Native.TryGetProcessInfo(hwnd, out var processName, out var processId, out var executablePath))
        {
            NotificationService.ShowWarning(notifyIcon, "Pencerenin işlem bilgisi okunamadı");
            return;
        }

        Console.WriteLine($"İşlem Adı: {processName}");
        Console.WriteLine($"İşlem ID: {processId}");
        if (executablePath != null) Console.WriteLine($"Yol: {executablePath}");

        if (!processName.ToLower().Contains("doganium"))
        {
            NotificationService.ShowWarning(notifyIcon, "Bu Doganium Sorgulama ekranı değil!");
            SystemSounds.Hand.Play();
            return;
        }

        NotificationService.ShowInfo(notifyIcon, "Doğanium penceresi yakalandı. Metin işleme başlatılıyor...");

        var uiText = UiAutomationTextReader.ReadText(hwnd);
        if (string.IsNullOrWhiteSpace(uiText))
        {
            Console.WriteLine("Erişilebilir metin bulunamadı");
            SystemSounds.Hand.Play();
            return;
        }

        Console.WriteLine("UI Automation ile alındı");
        Console.WriteLine($"Metin uzunluğu: {uiText.Length}");

        var formatted = DoganiumTextFormatter.Format(uiText, queryType, notifyIcon);
        if (string.IsNullOrEmpty(formatted))
        {
            SystemSounds.Hand.Play();
            return;
        }

        Clipboard.SetText(formatted);
        SystemSounds.Asterisk.Play();
    }
}