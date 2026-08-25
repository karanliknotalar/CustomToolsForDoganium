using System;
using System.Globalization;

namespace CustomToolsForDoganium.Capture
{
    /// <summary>Doğanium ekranından UI Automation ile alınan satır dizilerinden etikete göre
    /// (bir önceki/sonraki satır) alan değeri çıkarma ve biçimlendirme yardımcıları.
    /// Ekranda beklenen bir etiket bulunamazsa (o sorgu/müşteri tipinde alan hiç yoksa) hiçbir
    /// metot exception fırlatmaz, boş string döner — tek bir eksik alan yüzünden tüm yakalama
    /// işleminin çökmesi engellenir; formatlayıcı o alanı sadece boş bırakır. Anahtar bulunamadığı
    /// her durum ayrıca konsola loglanır, böylece hangi etiketin o ekranda hiç gelmediğini
    /// görebilirsin.</summary>
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

        /// <summary>Verilen anahtarı içeren satırdan hemen SONRAKİ satırı döner. Anahtar
        /// bulunamazsa veya dizinin son satırındaysa (sonrası olmadığı için) boş string döner
        /// ve durum konsola loglanır.</summary>
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

        /// <summary>Verilen anahtarı içeren satırdan hemen ÖNCEKİ satırı döner. Anahtar
        /// bulunamazsa veya dizinin ilk satırındaysa (öncesi olmadığı için) boş string döner
        /// ve durum konsola loglanır.</summary>
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