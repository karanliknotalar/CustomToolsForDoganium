using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace CustomToolsForDoganium
{
    public abstract class ForExel
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        private static Thread _workerThread;
        private static volatile bool _isRunning;

        public static void Start()
        {
            if (_isRunning) return;
            _isRunning = true;

            _workerThread = new Thread(() =>
            {
                while (_isRunning)
                {
                    try
                    {
                        if (IsExcelActive())
                            CheckAndConvertClipboard();
                    }
                    catch
                    {
                        Console.Error.WriteLine("Error while trying to get clipboard");
                    }

                    Thread.Sleep(1000);
                }
            });

            _workerThread.SetApartmentState(ApartmentState.STA);
            // _workerThread.IsBackground = true; // Ana program kapanınca bu da kapanır
            _workerThread.Start();
        }

        public static void Stop()
        {
            _isRunning = false;
        }

        private static bool IsExcelActive()
        {
            var hWnd = GetForegroundWindow();
            if (hWnd == IntPtr.Zero) return false;
            var sb = new StringBuilder(256);
            return GetWindowText(hWnd, sb, 256) > 0 && sb.ToString().Contains("Excel");
        }

        private static string _lastProcessedText = "";
        private static void CheckAndConvertClipboard()
        {
            if (!IsExcelActive()) return;
            var rawText = GetClipboardText();
            if (string.IsNullOrEmpty(rawText) || rawText == _lastProcessedText) return;
            if (!rawText.Contains("TC:") && !rawText.Contains("Vergi:")) return;

            var formattedText = ParseAndFormat(rawText);
            if (string.IsNullOrEmpty(formattedText)) return;

            _lastProcessedText = formattedText;
            
            SetClipboardText(formattedText);

            Console.WriteLine("[EXEL] Kopyalanan veri exel formatına uyarlandı.");
            Console.WriteLine($"[EXEL] Uyarlanan Veri: {formattedText}");
        }

        
        [DllImport("user32.dll")]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        private static extern bool CloseClipboard();

        private static bool IsClipboardBusy()
        {
            if (!OpenClipboard(IntPtr.Zero))
                return true;

            CloseClipboard();
            return false;
        }
        
        private static string GetClipboardText()
        {
            try
            {
                return Clipboard.GetText();
            }
            catch (Exception e)
            {
                Console.WriteLine("[EXEL] PANO OKUMA HATASI: " + e.Message + e);
                return null;
            }
        }

        private static void SetClipboardText(string text)
        {
            if (!IsExcelActive()) return;
            try
            {
                Clipboard.SetText(text);
            }
            catch (Exception e)
            {
                Console.WriteLine("[EXEL] PANO YAZMA HATASI:" + e.Message);
            }
        }

        private static string ParseAndFormat(string input)
        {
            try
            {
                var lines = input.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length < 2) return null;

                var sb = new StringBuilder();

                if (lines[1].Contains("TC:"))
                {
                    var adSoyad = lines[0].Trim();
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
                    var firmaAdi = lines[0].Trim();
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