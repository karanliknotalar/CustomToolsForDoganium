using System;
using System.Diagnostics;
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
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

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

                    Thread.Sleep(100);
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

            GetWindowThreadProcessId(hWnd, out var processId);

            try
            {
                var p = Process.GetProcessById((int)processId);
                return p.ProcessName.ToUpper() == "EXCEL";
            }
            catch
            {
                return false;
            }
        }


        private static void CheckAndConvertClipboard()
        {
            var rawText = GetClipboardText();
            if (string.IsNullOrEmpty(rawText)) return;
            if (!rawText.Contains("TC:") && !rawText.Contains("Vergi:")) return;

            var formattedText = ParseAndFormat(rawText);
            if (string.IsNullOrEmpty(formattedText)) return;

            SetClipboardText(formattedText);

            Console.WriteLine("[EXEL] Kopyalanan veri exel formatına uyarlandı.");
            Console.WriteLine($"[EXEL] Uyarlanan Veri: {formattedText}");
        }

        private static string GetClipboardText()
        {
            for (var i = 0; i < 3; i++)
            {
                try
                {
                    return Clipboard.GetText();
                }
                catch
                {
                    Thread.Sleep(100);
                }
            }

            return null;
        }

        private static void SetClipboardText(string text)
        {
            for (var i = 0; i < 3; i++)
            {
                try
                {
                    Clipboard.SetText(text);
                    return;
                }
                catch
                {
                    Thread.Sleep(100);
                }
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
            catch
            {
                Console.WriteLine("Hatalı işlem");
                return null;
            }
        }

        private static string GetSafeValue(string[] lines, int index, string prefix)
        {
            return index >= lines.Length ? "" : lines[index].Replace(prefix, "").Trim();
        }
    }
}