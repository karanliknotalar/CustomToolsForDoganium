using System.Globalization;

namespace CustomToolsForDoganium.Capture;

/// <summary>
/// Doğanium'dan gelen ham sigortalı/müşteri adını, bilinen bir ifadeyi içeriyorsa
/// tamamen o ifadeye karşılık gelen temiz isimle değiştirir (kısmi/parça değiştirme değil —
/// "içeriyorsa" burada bir TETİKLEYİCİ, eşleşince tüm isim ReplaceWith değeriyle değişir).
/// Liste sırayla taranır, İLK eşleşen kazanır. Yeni bir eşleşme eklemek istediğinde
/// <see cref="Replacements"/> listesine bir satır eklemen yeterli.
/// Eşleştirme <c>tr-TR</c> kültürüne göre büyük/küçük harf duyarsızdır — Türkçe'ye özgü
/// İ/i ve I/ı (noktalı/noktasız I) farkı da doğru şekilde göz ardı edilir.
/// </summary>
internal static class CustomerNameNormalizer
{
    private static readonly CompareInfo TurkishCompareInfo = CultureInfo.GetCultureInfo("tr-TR").CompareInfo;

    private static readonly (string Contains, string ReplaceWith)[] Replacements =
    [
        ("Nuroğlu Oto Galerisi", "Nuroğlu Oto Galeri"),
        ("Tr Kale Yapı İnşaat", "TrKale Yapı İnşaat"),
        ("GÜRSOYPLUS OTOMOTİV", "Gürsoyplus Otomotiv"),
        ("GÜRSOY NAKLİYAT", "Gürsoy Nakliyat"),
        ("MELSEM OTOMOTİV TURİZM", "Melsem Otomotiv"),
        ("TAYYAR MADEN İŞLETMELERİ", "Tayyar Maden İşletmeleri"),
        ("AZEEM PRİVE TEKSTİL", "Azeem Prive Tekstil"),
        ("SAMOUR TOUR İNŞAAT", "Samour Turizm"),
        ("ARSENDE GIDA İNŞAAT", "Arsende Gıda İnşaat"),
        ("ELEZOĞLU PETROL ÜRÜNLERİ", "Elezoğlu Petrol Ürünleri"),
        ("TRANS BEKTAŞ LOJİSTİK", "Trans Bektaş Lojistik"),
        ("YILMAZLAR PETROL NAKLİYAT", "Yılmazlar Petrol Nakliyat"),
        ("BÜLBÜLOĞLU İNŞAAT", "Bülbüloğlu İnşaat"),
        ("LEADERS VIP GROUP", "Leaders Vip Group"),
        ("EROĞLU ER PETROL", "Eroğlu Er Petrol"),
        ("SZK İNŞAAT TİCARET", "SZK İnşaat Ticaret"),
        ("SÜMELA NAKLİYAT HARFİYAT", "Sümela Nakliyat Harfiyat"),
        ("AKSU 34 NAKLİYE", "Aksu 34 Nakliye Otomotiv"),
        ("USTAOĞLU TURİZM NAKLİYAT", "Ustaoğlu Turizm Nakliyat"),
        ("HARMAN OTO KAFE", "Harman Oto Kafe"),
        ("AYH LOJİSTİK OTOMOTİV", "Ayh Lojistik Otomotiv"),
    ];

    internal static string Apply(string customerName)
    {
        if (string.IsNullOrEmpty(customerName)) return customerName;

        foreach (var (contains, replaceWith) in Replacements)
        {
            if (TurkishCompareInfo.IndexOf(customerName, contains, CompareOptions.IgnoreCase) >= 0)
                return replaceWith;
        }

        return customerName;
    }
}