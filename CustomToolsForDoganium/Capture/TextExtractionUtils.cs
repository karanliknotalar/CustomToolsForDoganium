using System;
using System.Globalization;

namespace CustomToolsForDoganium.Capture;

/// <summary>Doğanium ekranından UI Automation ile alınan satır dizilerinden etikete göre
/// (bir önceki/sonraki satır) alan değeri çıkarma ve biçimlendirme yardımcıları.</summary>
internal static class TextExtractionUtils
{
    internal static string GetDate(string text)
    {
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
        var tr = new CultureInfo("tr-TR");
        return tr.TextInfo.ToTitleCase(text.Trim().ToLower(tr));
    }

    internal static string GetVehicleGroupName(string text)
    {
        return GetToTitleCase(text.Split(' ')[2]);
    }

    internal static string GetNextValue(string[] arr, string key)
    {
        if (arr == null || arr.Length == 0) return null;

        for (var i = 0; i < arr.Length - 1; i++)
        {
            if (arr[i].Contains(key)) return arr[i + 1];
        }

        return null;
    }

    internal static string GetBeforeValue(string[] arr, string key)
    {
        if (arr == null || arr.Length == 0) return null;

        for (var i = 0; i < arr.Length - 1; i++)
        {
            if (arr[i].Contains(key)) return arr[i - 1];
        }

        return null;
    }
}