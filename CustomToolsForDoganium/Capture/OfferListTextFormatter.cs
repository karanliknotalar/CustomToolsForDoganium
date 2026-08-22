using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CustomToolsForDoganium.Notifications;

namespace CustomToolsForDoganium.Capture
{
    /// <summary>"Yeniden Sorgula" ekranındaki teklif satırlarını fiyata göre sıralı özet metne çevirir.</summary>
    internal static class OfferListTextFormatter
    {
        internal static string Format(string[] lines, InsuranceQueryType queryType, NotifyIcon notifyIcon)
        {
            var rowLines = lines
                .Where(line => line.Trim().StartsWith(";;", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (rowLines.Count == 0) return null;

            var offers = new List<InsuranceOffer>();
            var counter = 0;

            foreach (var line in rowLines)
            {
                var parts = line.Split(';');
                if (parts.Length < 7) continue;

                var companyName = parts[2].Trim();
                var priceStr = parts[(int)queryType].Trim();

                priceStr = priceStr.Split(',')[0].Split('.')[0];
                if (!int.TryParse(priceStr, out var price)) continue;

                var offerNumber = "";
                var index = (int)queryType + 3;
                var offerSection = (index < parts.Length ? parts[index] : null)?.Trim() ?? string.Empty;
                var companyUpper = companyName.ToUpper();

                if (!string.IsNullOrWhiteSpace(offerSection) && !companyUpper.Contains("HDI") &&
                    !companyUpper.Contains("SOMPO") && !companyUpper.Contains("HEPİYİ") &&
                    !companyUpper.Contains("AXA") && counter < 6)
                {
                    var match = Regex.Match(offerSection, @"[\d/]{7,}");
                    if (match.Success) offerNumber = match.Value.Trim();
                }

                if (offerSection.Contains("Bilgi") || offerSection.Contains("Hata"))
                {
                    companyName += " (?)";
                }

                offers.Add(new InsuranceOffer
                {
                    CompanyName = companyName.ToUpper(),
                    Price = price,
                    OfferNumber = offerNumber
                });
                counter++;
            }

            offers = offers.OrderBy(o => o.Price).ToList();

            var result = new StringBuilder();
            foreach (var offer in offers)
            {
                result.AppendLine(string.IsNullOrWhiteSpace(offer.OfferNumber)
                    ? $"{offer.Price} TL {offer.CompanyName}"
                    : $"{offer.Price} TL {offer.CompanyName} - {offer.OfferNumber}");
            }

            NotificationService.ShowInfo(notifyIcon, "Doganium Sorgu ekranından veriler kopyalandı!");
            return result.ToString();
        }
    }
}
