using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Windows.Forms;

namespace CustomToolsForDoganium.Startup
{
    /// <summary>Uygulamanın yönetici (administrator) yetkisiyle çalışıp çalışmadığını kontrol eder
    /// ve gerekirse yükseltilmiş yetkiyle yeniden başlatır.</summary>
    internal static class AdminElevationService
    {
        internal static bool IsRunningAsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        internal static void RestartAsAdministrator()
        {
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = true,
                WorkingDirectory = Environment.CurrentDirectory,
                FileName = Application.ExecutablePath,
                Verb = "runas"
            };
            Process.Start(startInfo);
        }
    }
}
