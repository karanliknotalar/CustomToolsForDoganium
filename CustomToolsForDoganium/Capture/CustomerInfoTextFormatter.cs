using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CustomToolsForDoganium.Notifications;

namespace CustomToolsForDoganium.Capture;

/// <summary>"Genel Bilgiler" ekranındaki müşteri ve araç bilgilerini etiketli özet metne çevirir.</summary>
internal static class CustomerInfoTextFormatter
{
    private static readonly Regex DocumentSeriesRegex = new(@"^[A-Za-z]{1,2}\d{6}$", RegexOptions.Compiled);

    internal static string Format(string[] lines, NotifyIcon notifyIcon)
    {
        if (lines.Length == 0) return null;

        var vehicleGroupName =
            TextExtractionUtils.GetVehicleGroupName(TextExtractionUtils.GetBeforeValue(lines, "Kasko Değeri"));
        var vehicleTypeCode = TextExtractionUtils.GetNextValue(lines, "Tip Kodu").Split(' ').First();
        var vehicleBrandCode = TextExtractionUtils.GetBeforeValue(lines, "Şasi No", 3).Split(' ').First();
        var vehicleModelYear = TextExtractionUtils.GetNextValue(lines, "Model Yılı");
        var documentSeries = TextExtractionUtils.GetNextValue(lines, "Plaka");

        if (!DocumentSeriesRegex.IsMatch(documentSeries))
        {
            documentSeries = "";
        }

        var motorNo = TextExtractionUtils.GetBeforeValue(lines, "Tescil Tarihi");
        var chassisNo = TextExtractionUtils.GetNextValue(lines, "Şasi No");

        var birthDate = TextExtractionUtils.GetDate(TextExtractionUtils.GetBeforeValue(lines, "TC / Vergi No"));
        var identifyNo = TextExtractionUtils.GetNextValue(lines, "TC / Vergi No");
        var plate = TextExtractionUtils.GetBeforeValue(lines, "Plakanın Son Sorgusunu Getir").Replace(" ", "");
        var customerName = CustomerNameNormalizer.Apply(
            TextExtractionUtils.GetToTitleCase(TextExtractionUtils.GetBeforeValue(lines, "gcInsuranceInformation")));

        var result = new StringBuilder();
        result.AppendLine($"Sigortalı: {customerName}");

        switch (identifyNo.Length)
        {
            case 11:
                result.AppendLine($"TC: {identifyNo}")
                    .AppendLine($"DT: {birthDate}");
                break;
            case 10:
                result.AppendLine($"Vergi: {identifyNo}");
                break;
        }

        result.AppendLine($"Plaka: {plate}")
            .AppendLine($"Seri: {documentSeries}")
            .AppendLine($"Motor: {motorNo}")
            .AppendLine($"Şasi: {chassisNo}")
            .AppendLine($"Kullanım Şekli: {vehicleGroupName}")
            .AppendLine($"Model Yılı: {vehicleModelYear}")
            .AppendLine($"Marka Kodu: {vehicleBrandCode}{vehicleTypeCode}");

        NotificationService.ShowInfo(notifyIcon, "Doganium müşteri ve araç bilgileri ekranından veriler kopyalandı!");
        return result.ToString();
    }
}