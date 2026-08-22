using System;
using System.Threading;
using System.Windows.Forms;
using CustomToolsForDoganium.Capture;
using CustomToolsForDoganium.Hotkeys;
using CustomToolsForDoganium.Native;
using CustomToolsForDoganium.Notifications;

namespace CustomToolsForDoganium.App
{
    /// <summary>Uygulamanın sonsuz döngüsünü çalıştırır: konsol penceresini gizler,
    /// yakalama kısayolunu ve genel kısayol yöneticisini her tick'te kontrol eder.</summary>
    internal sealed class ApplicationRunner(
        IntPtr consoleWindow,
        NotifyIcon notifyIcon,
        HotkeyActionManager hotkeyManager,
        CaptureHotkeyWatcher captureHotkeyWatcher,
        DoganiumCaptureService captureService)
    {
        internal void Run()
        {
            while (true)
            {
                try
                {
                    if (ConsoleWindowNative.IsMinimized(consoleWindow))
                        ConsoleWindowNative.Hide(consoleWindow);

                    if (captureHotkeyWatcher.TryGetPressedQueryType(out var queryType))
                    {
                        captureService.Capture(queryType);
                        Thread.Sleep(800);
                    }

                    hotkeyManager.Poll();

                    Application.DoEvents();
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    Console.WriteLine("Stack Trace: " + ex.StackTrace);
                    NotificationService.ShowWarning(notifyIcon,
                        "TEKRAR DENE! UYGULAMA EKRANI YAKALARKEN BİR HATA OLUŞTU.");
                }
            }
        }
    }
}
