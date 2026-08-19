using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace CustomToolsForDoganium
{
    internal static class ClipboardInsert
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_CONTROL = 0x11;
        private const int VK_Q = 0x51;

        private static bool _wasQDown;

        public static bool IsHotkeyPressed()
        {
            var ctrlDown = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
            var qDown = (GetAsyncKeyState(VK_Q) & 0x8000) != 0;

            var triggered = ctrlDown && qDown && !_wasQDown;
            _wasQDown = qDown;

            return triggered;
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

            var processName = proc.ProcessName.ToLower();

            var isNotepadPlusPlus = processName.Contains("notepad++");
            var isExcel = processName.Contains("excel");

            // Sadece Notepad++ veya Excel önplandaysa çalış
            if (!isNotepadPlusPlus && !isExcel)
            {
                return;
            }

            if (!Clipboard.ContainsText())
            {
                Tools.ShowNotifyWarning(notifyIcon, "Panoda metin bulunamadı!");
                return;
            }

            var previousClipboard = Clipboard.GetText();

            var summary = isNotepadPlusPlus
                ? BuildSummaryForNotepad(previousClipboard)
                : BuildSummaryForExel(previousClipboard);

            if (string.IsNullOrEmpty(summary))
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

            //Tools.ShowNotifyInfo(notifyIcon, "Özet bilgi " + (isNotepadPlusPlus ? "Notepad++" : "Excel") + " içine yapıştırıldı!");
            Console.WriteLine("[" + (isNotepadPlusPlus ? "Notepad++" : "Excel") + "]: " + summary);
        }

        private static string BuildSummaryForNotepad(string text)
        {
            var lines = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

            var sigortaliLine = lines.FirstOrDefault(l =>
                l.Trim().StartsWith("Sigortalı:", StringComparison.OrdinalIgnoreCase));
            var plakaLine = lines.FirstOrDefault(l =>
                l.Trim().StartsWith("Plaka:", StringComparison.OrdinalIgnoreCase));

            if (sigortaliLine == null) return null;

            var policeLine = lines.FirstOrDefault(l =>
                Regex.IsMatch(l.Trim(), @"^(Trafik|Kasko|Imm|TSS|Konut|İşyeri)#+", RegexOptions.IgnoreCase));

            var sigortali = sigortaliLine.Split([':'], 2)[1].Trim();

            var policeTuru = policeLine != null
                ? Regex.Match(policeLine.Trim(), @"^(Trafik|Kasko|Imm|TSS|Konut|İşyeri)", RegexOptions.IgnoreCase).Value
                : "POLİÇETÜRÜ";

            var tarih = DateTime.Now.ToString("dd.MM.yyyy");

            if (plakaLine == null)
            {
                return $"{tarih} - {sigortali} - {policeTuru}";
            }

            var plaka = plakaLine.Split([':'], 2)[1].Trim();

            return $"{tarih} - {sigortali} - {plaka} - {policeTuru}";
        }

        private static string BuildSummaryForExel(string input)
        {
            try
            {
                var lines = input.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length < 2) return null;

                var sb = new StringBuilder();

                if (lines[1].Contains("TC:"))
                {
                    var adSoyad = GetSafeValue(lines, 0, "Sigortalı:");
                    var tc = GetSafeValue(lines, 1, "TC:");
                    var dt = GetSafeValue(lines, 2, "DT:");
                    var plaka = GetSafeValue(lines, 3, "Plaka:");
                    var seri = GetSafeValue(lines, 4, "Seri:");

                    sb.Append(adSoyad);
                    sb.Append("\t\t");
                    sb.Append(tc);
                    sb.Append("\t");
                    sb.Append(dt);
                    sb.Append("\t");
                    sb.Append(plaka);
                    sb.Append("\t");
                    sb.Append(seri);
                    sb.Append("\t");
                }
                else if (lines[1].Contains("Vergi:"))
                {
                    var firmaAdi = GetSafeValue(lines, 0, "Sigortalı:");
                    var vergi = GetSafeValue(lines, 1, "Vergi:");
                    var plaka = GetSafeValue(lines, 2, "Plaka:");
                    var seri = GetSafeValue(lines, 3, "Seri:");

                    sb.Append(firmaAdi);
                    sb.Append("\t\t");
                    sb.Append(vergi);
                    sb.Append("\t\t");
                    sb.Append(plaka);
                    sb.Append("\t");
                    sb.Append(seri);
                    sb.Append("\t");
                }
                else
                {
                    return null;
                }

                return sb.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine("[EXEL]  Hatalı işlem: " + e.Message);
                return null;
            }
        }

        private static string GetSafeValue(string[] lines, int index, string prefix)
        {
            return index >= lines.Length ? "" : lines[index].Replace(prefix, "").Trim();
        }
    }
}