namespace CustomToolsForDoganium.Capture;

/// <summary>Doğanium "Yeniden Sorgula" ekranındaki teklif satırında hangi sütunun (fiyat) okunacağını belirler.
/// Sayısal değerler, ham verideki noktalı virgülle ayrılmış sütun index'ine karşılık gelir.</summary>
internal enum InsuranceQueryType
{
    None = 0,
    Trafik = 4,
    Kasko = 5,
    Imm = 6
}