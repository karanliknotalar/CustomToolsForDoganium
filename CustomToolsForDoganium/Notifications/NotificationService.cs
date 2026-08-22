using System;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CustomToolsForDoganium.Notifications
{
    /// <summary>Tray balon bildirimlerini gösterir ve Windows bildirim ayarlarının açık olmasını sağlar.</summary>
    internal static class NotificationService
    {
        private const string Title = "";

        internal static void ShowInfo(NotifyIcon notifyIcon, string message)
        {
            notifyIcon.ShowBalloonTip(1000, Title, message, ToolTipIcon.Info);
            Console.WriteLine(message);
        }

        internal static void ShowWarning(NotifyIcon notifyIcon, string message)
        {
            notifyIcon.ShowBalloonTip(1000, Title, message, ToolTipIcon.Warning);
            Console.WriteLine(message);
        }

        internal static void EnsureNotificationsEnabled()
        {
            try
            {
                var appName = Application.ExecutablePath;
                var regPath = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\{appName}";

                using var key = Registry.CurrentUser.CreateSubKey(regPath);
                key?.SetValue("Enabled", 1, RegistryValueKind.DWord);
                key?.SetValue("ShowInActionCenter", 0, RegistryValueKind.DWord);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Bildirim ayarı yapılamadı: " + ex.Message);
            }
        }
    }
}
