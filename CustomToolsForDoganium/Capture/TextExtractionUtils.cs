using System;
using System.Globalization;

namespace CustomToolsForDoganium.Capture
{
    internal static class TextExtractionUtils
    {
        internal static string GetDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            return DateTime.TryParseExact(
                text.Trim(),
                "dd MMMM yyyy dddd",
                new CultureInfo("tr-TR"),
                DateTimeStyles.None,
                out var date)
                ? date.ToString("dd.MM.yyyy")
                : string.Empty;
        }

        internal static string GetToTitleCase(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            var tr = new CultureInfo("tr-TR");
            return tr.TextInfo.ToTitleCase(text.Trim().ToLower(tr));
        }

        internal static string GetVehicleGroupName(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            var parts = text.Split(' ');
            return parts.Length > 2 ? GetToTitleCase(parts[2]) : string.Empty;
        }
        
        internal static string GetNextValue(string[] arr, string key, int step = 0)
        {
            if (arr == null)
            {
                LogKeyNotFound(nameof(GetNextValue), key, "dizi null");
                return string.Empty;
            }

            for (var i = 0; i < arr.Length - 1; i++)
            {
                if (arr[i].Contains(key)) return arr[i + (1 + step)];
            }

            LogKeyNotFound(nameof(GetNextValue), key, "ekranda bu etiket bulunamadı");
            return string.Empty;
        }
        
        internal static string GetBeforeValue(string[] arr, string key, int step = 0)
        {
            if (arr == null)
            {
                LogKeyNotFound(nameof(GetBeforeValue), key, "dizi null");
                return string.Empty;
            }

            for (var i = 1; i < arr.Length; i++)
            {
                if (arr[i].Contains(key)) return arr[i - (1 + step)];
            }

            LogKeyNotFound(nameof(GetBeforeValue), key, "ekranda bu etiket bulunamadı");
            return string.Empty;
        }

        private static void LogKeyNotFound(string method, string key, string reason)
        {
            Console.WriteLine($"[TextExtractionUtils.{method}] Anahtar bulunamadı: \"{key}\" ({reason})");
        }
    }
}