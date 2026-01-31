using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;

namespace CustomToolsForDoganium
{
    internal static partial class Program
    {
        // ===================== WIN32 API =====================
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(Keys vKey);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // Pencere gizleme ve durum kontrolü için gerekli API'ler
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        private const int SW_HIDE = 0;
        private const int SW_RESTORE = 9;

        private static NotifyIcon _notifyIcon;
        private static IntPtr _consoleWindow;

        // ===================== MAIN =====================
        [STAThread]
        private static void Main()
        {
            _consoleWindow = GetConsoleWindow();

            if (!IsRunAsAdministrator())
            {
                Console.WriteLine("⚠️  Yönetici yetkisi gerekli!");
                Console.WriteLine();
                Console.WriteLine("[1] Otomatik olarak yönetici yetkisiyle yeniden başlat");
                Console.WriteLine("[2] Çıkış (manuel olarak yönetici yetkisiyle çalıştırın)");
                Console.WriteLine();
                Console.Write("Seçiminiz (1/2): ");

                var choice = Console.ReadLine();

                if (choice != "1") return;
                Console.WriteLine("Yeniden başlatılıyor...");
                RestartAsAdministrator();

                return;
            }

            _notifyIcon = new NotifyIcon
            {
                Icon = SystemIcons.Shield,
                Text = "Doganium Araçları",
                Visible = true
            };

            _notifyIcon.DoubleClick += (s, e) => { ShowWindow(_consoleWindow, SW_RESTORE); };

            _notifyIcon.ContextMenu = new ContextMenu(new[]
            {
                new MenuItem("Göster", (s, e) => ShowWindow(_consoleWindow, SW_RESTORE)),
                new MenuItem("-"),
                new MenuItem("Çıkış", (s, e) =>
                {
                    _notifyIcon.Visible = false;
                    Environment.Exit(0);
                })
            });

            Console.WriteLine("✅ Yönetici yetkisiyle çalışıyor");
            Console.WriteLine("Ctrl + Shift + C : Erişilebilir metni yakala");
            Console.WriteLine("💡 Pencereyi küçülttüğünde saat yanına (tray) gizlenecektir.");

            while (true)
            {
                try
                {
                    if (IsIconic(_consoleWindow))
                    {
                        ShowWindow(_consoleWindow, SW_HIDE);
                    }

                    if (IsHotkeyPressed())
                    {
                        CaptureText_SafeWindowsOnly();
                        Thread.Sleep(800);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Hata: " + ex.Message);
                }

                Application.DoEvents();
                Thread.Sleep(100);
            }
        }

        private static bool IsRunAsAdministrator()
        {
            try
            {
                var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private static void RestartAsAdministrator()
        {
            try
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
            catch (Exception ex)
            {
                Console.WriteLine($"Hata: {ex.Message}");
                Console.ReadKey();
            }
        }

        private static bool IsHotkeyPressed()
        {
            return (GetAsyncKeyState(Keys.ControlKey) & 0x8000) != 0 &&
                   (GetAsyncKeyState(Keys.ShiftKey) & 0x8000) != 0 &&
                   (GetAsyncKeyState(Keys.C) & 0x8000) != 0;
        }

        // ===================== SMART CAPTURE =====================
        private static void CaptureText_SafeWindowsOnly()
        {
            var hwnd = GetForegroundWindow();
            Console.WriteLine($"\n=== YENİ YAKALAMA DENEMESİ ===");
            Console.WriteLine($"HWND: {hwnd}");

            if (hwnd == IntPtr.Zero)
            {
                Tools.ShowNotifyWarning(_notifyIcon, "Aktif pencere bulunamadı");
                return;
            }

            var length = GetWindowTextLength(hwnd);
            var windowTitle = new StringBuilder(length + 1);

            GetWindowText(hwnd, windowTitle, windowTitle.Capacity);
            
            if (!windowTitle.ToString().Contains("Sorgulama Ekranı"))
            {
                Tools.ShowNotifyWarning(_notifyIcon, "Aktif pencere Dogaium sorgulama ekranı değil!");
                return;
            }
            Console.WriteLine($"Pencere Başlığı: '{windowTitle}'");

            GetWindowThreadProcessId(hwnd, out var processId);

            try
            {
                var proc = Process.GetProcessById((int)processId);
                Console.WriteLine($"İşlem Adı: {proc.ProcessName}");
                Console.WriteLine($"İşlem ID: {processId}");
                if (proc.MainModule != null) Console.WriteLine($"Yol: {proc.MainModule.FileName}");

                if (!proc.ProcessName.ToLower().Contains("doganium"))
                {
                    Tools.ShowNotifyWarning(_notifyIcon, "Bu Doganium Sorgulama ekranı değil!");
                    System.Media.SystemSounds.Hand.Play();
                    return;
                }
            }
            catch (Exception ex)
            {
                Tools.ShowNotifyWarning(_notifyIcon, $"İşlem bilgisi alınamadı: {ex.Message}" +
                                                     $"\r\n(Muhtemelen yetki sorunu - uygulamayı yönetici olarak çalıştırın)");
                System.Media.SystemSounds.Hand.Play();
                return;
            }

            Tools.ShowNotifyInfo(_notifyIcon, "Doğanium penceresi yakalandı. Metin işleme başlatılıyor...");

            var uiText = ReadUiAutomationText(hwnd);
            if (!string.IsNullOrWhiteSpace(uiText))
            {
                Console.WriteLine("UI Automation ile alındı");
                Console.WriteLine($"Metin uzunluğu: {uiText.Length}");

                var result = ProcessAndFormatText(uiText);
                Clipboard.SetText(uiText);
                Clipboard.SetText(result);
                Console.WriteLine(result);

                System.Media.SystemSounds.Asterisk.Play();
                return;
            }

            Console.WriteLine("Erişilebilir metin bulunamadı");
            System.Media.SystemSounds.Hand.Play();
        }

        private static string ProcessAndFormatText(string rawText)
        {
            try
            {
                var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var rowLines = lines.Where(line => line.Trim().StartsWith(";;", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (rowLines.Count == 0) return null;

                var offers = new List<InsuranceOffer>();
                foreach (var line in rowLines)
                {
                    var parts = line.Split(';');
                    if (parts.Length < 6) continue;

                    var companyName = parts[2].Trim();
                    var priceStr = parts[4].Trim();

                    priceStr = priceStr.Split(',')[0].Split('.')[0];
                    if (!int.TryParse(priceStr, out var price)) continue;

                    var offerNumber = "";
                    if (parts.Length > 7)
                    {
                        var offerSection = parts[7].Trim();
                        var companyUpper = companyName.ToUpper();
                        if (!string.IsNullOrWhiteSpace(offerSection) && !companyUpper.Contains("HDI") &&
                            !companyUpper.Contains("SOMPO") && !companyUpper.Contains("HEPİYİ"))
                        {
                            var match = Regex.Match(offerSection, @"[\d/]+");
                            if (match.Success) offerNumber = match.Value.Trim();
                        }
                    }

                    offers.Add(new InsuranceOffer
                        { CompanyName = companyName.ToUpper(), Price = price, OfferNumber = offerNumber });
                }

                offers = offers.OrderBy(o => o.Price).ToList();
                var result = new StringBuilder();
                foreach (var offer in offers)
                {
                    result.AppendLine(string.IsNullOrWhiteSpace(offer.OfferNumber)
                        ? $"{offer.Price} TL {offer.CompanyName}"
                        : $"{offer.Price} TL {offer.CompanyName} - {offer.OfferNumber}");
                }

                Tools.ShowNotifyInfo(_notifyIcon, "Doganium Sorgu ekranından veriler kopyalandı!");
                return result.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static string ReadUiAutomationText(IntPtr hwnd)
        {
            try
            {
                var root = AutomationElement.FromHandle(hwnd);
                if (root == null) return null;
                var sb = new StringBuilder();
                WalkUiAutomation(root, sb, 0);
                return sb.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static void WalkUiAutomation(AutomationElement el, StringBuilder sb, int depth)
        {
            try
            {
                if (depth > 7) return;
                if (!string.IsNullOrWhiteSpace(el.Current.Name)) sb.AppendLine(el.Current.Name);

                if (el.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
                    sb.AppendLine(((ValuePattern)valuePattern).Current.Value);

                if (el.TryGetCurrentPattern(TextPattern.Pattern, out var textPattern))
                    sb.AppendLine(((TextPattern)textPattern).DocumentRange.GetText(-1));

                var children = el.FindAll(TreeScope.Children, Condition.TrueCondition);
                foreach (AutomationElement c in children) WalkUiAutomation(c, sb, depth + 1);
            }
            catch
            {
                // ignored
            }
        }
    }
}