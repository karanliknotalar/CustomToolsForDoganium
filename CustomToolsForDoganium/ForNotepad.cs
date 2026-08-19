using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace CustomToolsForDoganium
{
    internal static class ForNotepad
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(Keys vKey);

        public static bool IsHotkeyPressed()
        {
            return (GetAsyncKeyState(Keys.ControlKey) & 0x8000) != 0 &&
                   (GetAsyncKeyState(Keys.Q) & 0x8000) != 0;
        }

        public static void TryInsertSummary(NotifyIcon notifyIcon)
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return;

            GetWindowThreadProcessId(hwnd, out var processId);

            Process proc;
            try
            {
                proc = Process.GetProcessById((int)processId);
            }
            catch
            {
                return;
            }

            // Sadece Notepad++ önplandaysa çalış
            if (!proc.ProcessName.ToLower().Contains("notepad++"))
            {
                return;
            }

            if (!Clipboard.ContainsText())
            {
                Tools.ShowNotifyWarning(notifyIcon, "Panoda metin bulunamadı!");
                return;
            }

            var previousClipboard = Clipboard.GetText();
            var summary = BuildSummary(previousClipboard);

            if (summary == null)
            {
                Tools.ShowNotifyWarning(notifyIcon,
                    "Panodaki veri istenen formatta değil! (Sigortalı / Plaka bilgisi bulunamadı)");
                return;
            }

            // Geçici olarak özet metni panoya koy, yapıştır, sonra eski panoyu geri yükle
            Clipboard.SetText(summary);
            SendKeys.SendWait("^v");
            Thread.Sleep(150);
            Clipboard.SetText(previousClipboard);

            Tools.ShowNotifyInfo(notifyIcon, "Özet bilgi Notepad++ içine yapıştırıldı!");
        }

        private static string BuildSummary(string text)
        {
            var lines = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

            var sigortaliLine = lines.FirstOrDefault(l =>
                l.Trim().StartsWith("Sigortalı:", StringComparison.OrdinalIgnoreCase));
            var plakaLine = lines.FirstOrDefault(l =>
                l.Trim().StartsWith("Plaka:", StringComparison.OrdinalIgnoreCase));

            if (sigortaliLine == null || plakaLine == null) return null;

            var policeLine = lines.FirstOrDefault(l =>
                Regex.IsMatch(l.Trim(), @"^(Trafik|Kasko|Imm|TSS|Konut|İşyeri)#+", RegexOptions.IgnoreCase));

            var sigortali = sigortaliLine.Split([':'], 2)[1].Trim();
            var plaka = plakaLine.Split([':'], 2)[1].Trim();

            var policeTuru = policeLine != null
                ? Regex.Match(policeLine.Trim(), @"^(Trafik|Kasko|Imm|TSS|Konut|İşyeri)", RegexOptions.IgnoreCase).Value
                : "POLİÇETÜRÜ";

            var tarih = DateTime.Now.ToString("dd.MM.yyyy");

            return $"{tarih} - {sigortali} - {plaka} - {policeTuru}";
        }
    }
}