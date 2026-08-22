using System;
using System.Windows.Forms;
using CustomToolsForDoganium.Notifications;

namespace CustomToolsForDoganium.Capture
{
    /// <summary>Ham UI Automation metnini içeriğine bakarak uygun formatlayıcıya (teklif listesi /
    /// müşteri bilgisi) yönlendirir. Yeni bir ekran türü desteklenecekse yeni bir formatter yazılıp
    /// buraya bir dal olarak eklenir; mevcut formatterlar değişmez (Open/Closed).</summary>
    internal static class DoganiumTextFormatter
    {
        internal static string Format(string rawText, InsuranceQueryType queryType, NotifyIcon notifyIcon)
        {
            var lines = rawText.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

            if (rawText.Contains("Yeniden Sorgula"))
                return OfferListTextFormatter.Format(lines, queryType, notifyIcon);

            if (rawText.Contains("Genel Bilgiler"))
                return CustomerInfoTextFormatter.Format(lines, notifyIcon);

            NotificationService.ShowWarning(notifyIcon, "Aktif Doğanium sekmesi müşteri bilgi veya sorgu ekranı değil!");
            return null;
        }
    }
}
