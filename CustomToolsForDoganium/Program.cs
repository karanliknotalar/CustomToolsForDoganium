using System;
using CustomToolsForDoganium.Actions;
using CustomToolsForDoganium.App;
using CustomToolsForDoganium.Capture;
using CustomToolsForDoganium.Hotkeys;
using CustomToolsForDoganium.Native;
using CustomToolsForDoganium.Notifications;
using CustomToolsForDoganium.Startup;

namespace CustomToolsForDoganium
{
    /// <summary>Uygulama giriş noktası. Sadece kurulum (composition root) yapar; iş mantığı
    /// ilgili sınıflarda yaşar. Yeni bir kısayol eklemek için aşağıdaki hotkeyManager listesine
    /// yeni bir IHotkeyAction eklemek yeterlidir.</summary>
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            if (!AdminElevationService.IsRunningAsAdministrator())
            {
                Console.WriteLine("⚠️  Yönetici yetkisi gerekli!");
                Console.WriteLine();
                Console.WriteLine("[1] Otomatik olarak yönetici yetkisiyle yeniden başlat");
                Console.WriteLine("[2] Çıkış (manuel olarak yönetici yetkisiyle çalıştırın)");
                Console.WriteLine();
                Console.Write("Seçiminiz (1/2): ");

                if (Console.ReadLine() != "1") return;

                Console.WriteLine("Yeniden başlatılıyor...");
                AdminElevationService.RestartAsAdministrator();
                return;
            }

            var consoleWindow = ConsoleWindowNative.GetHandle();
            var notifyIcon = TrayIconFactory.Create(consoleWindow);

            var hotkeyManager = new HotkeyActionManager(notifyIcon, [
                new SummaryInsertAction(),
                new EmptyTemplateInsertAction(),
                new EmptyTemplateInsertCorporateAction()
                // yeni kısayol eklemek istediğinde buraya ekleyeceksin
            ]);

            var captureHotkeyWatcher = new CaptureHotkeyWatcher();
            var captureService = new DoganiumCaptureService(notifyIcon);

            Console.WriteLine("✅ Yönetici yetkisiyle çalışıyor");
            Console.WriteLine("Ctrl + Shift + C / X / Z : Erişilebilir metni yakala");
            Console.WriteLine("💡 Pencereyi küçülttüğünde saat yanına (tray) gizlenecektir.");
            NotificationService.EnsureNotificationsEnabled();

            new ApplicationRunner(consoleWindow, notifyIcon, hotkeyManager, captureHotkeyWatcher, captureService)
                .Run();
        }
    }
}
