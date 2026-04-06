using System;
using System.Globalization;
using System.Windows.Forms;

namespace CustomToolsForDoganium
{
    public static class Tools
    {
        private const string Title = "";

        public static void ShowNotifyInfo(NotifyIcon n, string m)
        {
            n.ShowBalloonTip(1000, Title,
                m, ToolTipIcon.Info);
            Console.WriteLine(m);
        }

        public static void ShowNotifyWarning(NotifyIcon n, string m)
        {
            n.ShowBalloonTip(1000, Title,
                m, ToolTipIcon.Warning);
            Console.WriteLine(m);
        }

        public static string GetDate(string text)
        {
            return DateTime.TryParseExact(
                text.Trim(),
                "dd MMMM yyyy dddd",
                new CultureInfo("tr-TR"),
                DateTimeStyles.None,
                out var date) ? date.ToString("dd.MM.yyyy") :
                string.Empty;
        }


        public static string GetToTitleCase(string text)
        {
            var tr = new CultureInfo("tr-TR");
            return tr.TextInfo.ToTitleCase(text.Trim().ToLower(tr));
        }

        public static string GetVehicleGroupName(string text)
        {
            return GetToTitleCase(text.Split(' ')[2]);
        }
        
        public static string GetNextValue(string[] arr, string key)
        {
            if (arr == null || arr.Length == 0)
                return null;

            for (var i = 0; i < arr.Length - 1; i++)
            {
                if (arr[i].Contains(key))
                    return arr[i + 1];
            }

            return null;
        }
        
        public static string GetBeforeValue(string[] arr, string key)
        {
            if (arr == null || arr.Length == 0)
                return null;

            for (var i = 0; i < arr.Length - 1; i++)
            {
                if (arr[i].Contains(key))
                    return arr[i - 1];
            }

            return null;
        }
        
        public static void EnsureNotificationsEnabled()
        {
            try
            {
                var appName = Application.ExecutablePath;
                var regPath = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\{appName}";

                using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(regPath);
                key?.SetValue("Enabled", 1, Microsoft.Win32.RegistryValueKind.DWord);
                key?.SetValue("ShowInActionCenter", 0, Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Bildirim ayarı yapılamadı: " + ex.Message);
            }
        }
    }
}