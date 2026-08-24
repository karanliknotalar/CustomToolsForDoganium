using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CustomToolsForDoganium.Capture;
using CustomToolsForDoganium.Helper;
using CustomToolsForDoganium.Hotkeys;
using CustomToolsForDoganium.Notifications;

namespace CustomToolsForDoganium.Actions;

/// <summary>Ctrl+Q: panodaki poliçe/sigortalı bilgisinden tek satırlık özet üretip yapıştırır.</summary>
internal sealed class SummaryInsertAction : IHotkeyAction
{
    private static readonly Regex PolicyLineRegex = new(@"^(\p{L}+)#+", RegexOptions.Compiled);

    public Keys VirtualKey => Keys.Q;
    public TargetApp SupportedApps => TargetApp.NotepadPlusPlus | TargetApp.Excel;

    public void Execute(string clipboardText, TargetApp activeApp, NotifyIcon notifyIcon)
    {
        if (clipboardText == null)
        {
            NotificationService.ShowWarning(notifyIcon, "Panoda metin bulunamadı!");
            return;
        }

        var summary = activeApp == TargetApp.NotepadPlusPlus
            ? BuildSummaryForNotepad(clipboardText)
            : BuildSummaryForExcel(clipboardText);

        if (string.IsNullOrEmpty(summary))
        {
            NotificationService.ShowWarning(notifyIcon,
                "Panodaki veri istenen formatta değil! (Sigortalı / Plaka bilgisi bulunamadı)");
            return;
        }

        ClipboardPasteHelper.PasteAndRestore(summary);
        Console.WriteLine($"[{activeApp}]: {summary}");
    }

    private static string BuildSummaryForNotepad(string text)
    {
        var lines = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        var sigortaliLine = lines.FirstOrDefault(l =>
            l.Trim().StartsWith("Sigortalı:", StringComparison.OrdinalIgnoreCase));
        if (sigortaliLine == null) return null;

        var plakaLine = lines.FirstOrDefault(l =>
            l.Trim().StartsWith("Plaka:", StringComparison.OrdinalIgnoreCase));
        var policeLine = lines.FirstOrDefault(l => PolicyLineRegex.IsMatch(l.Trim()));

        var sigortali = sigortaliLine.Split([':'], 2)[1].Trim();

        var policeTuru = policeLine != null
            ? TextExtractionUtils.GetToTitleCase(PolicyLineRegex.Match(policeLine.Trim()).Groups[1].Value)
            : "POLİÇETÜRÜ";

        var tarih = DateTime.Now.ToString("dd.MM.yyyy");

        if (plakaLine == null)
            return $"{tarih} - {sigortali} - {policeTuru}";

        var plaka = plakaLine.Split([':'], 2)[1].Trim();
        return $"{tarih} - {sigortali} - {plaka} - {policeTuru}";
    }

    private static string BuildSummaryForExcel(string input)
    {
        try
        {
            var lines = input.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2) return null;

            var sb = new StringBuilder();
            var sigortaliLine = lines.FirstOrDefault(l =>
                l.Trim().StartsWith("Sigortalı:", StringComparison.OrdinalIgnoreCase));

            if (sigortaliLine != null)
            {
                var tcLine = lines.FirstOrDefault(l => l.Trim().StartsWith("TC:", StringComparison.OrdinalIgnoreCase));
                var vergiLine =
                    lines.FirstOrDefault(l => l.Trim().StartsWith("Vergi:", StringComparison.OrdinalIgnoreCase));

                if (tcLine != null) AppendBireysel(sb, lines);
                else if (vergiLine != null) AppendKurumsal(sb, lines);
                else return null;
            }
            else if (lines[1].Contains("TC:")) AppendBireyselByIndex(sb, lines);
            else if (lines[1].Contains("Vergi:")) AppendKurumselByIndex(sb, lines);
            else return null;

            return sb.ToString();
        }
        catch (Exception e)
        {
            Console.WriteLine("[EXCEL] Hatalı işlem: " + e.Message);
            return null;
        }
    }

    private static void AppendBireysel(StringBuilder sb, string[] lines) => sb
        .Append(GetSafeValueByLabel(lines, "Sigortalı:")).Append("\t\t")
        .Append(GetSafeValueByLabel(lines, "TC:")).Append('\t')
        .Append(GetSafeValueByLabel(lines, "DT:")).Append('\t')
        .Append(GetSafeValueByLabel(lines, "Plaka:")).Append('\t')
        .Append(GetSafeValueByLabel(lines, "Seri:")).Append("\t\t")
        .Append(GetSafeValueByLabel(lines, "Tel:"));

    private static void AppendKurumsal(StringBuilder sb, string[] lines) => sb
        .Append(GetSafeValueByLabel(lines, "Sigortalı:")).Append("\t\t")
        .Append(GetSafeValueByLabel(lines, "Vergi:")).Append("\t\t")
        .Append(GetSafeValueByLabel(lines, "Plaka:")).Append('\t')
        .Append(GetSafeValueByLabel(lines, "Seri:")).Append("\t\t")
        .Append(GetSafeValueByLabel(lines, "Tel:"));

    private static void AppendBireyselByIndex(StringBuilder sb, string[] lines) => sb
        .Append(GetSafeValue(lines, 0, "Sigortalı:")).Append("\t\t")
        .Append(GetSafeValue(lines, 1, "TC:")).Append('\t')
        .Append(GetSafeValue(lines, 2, "DT:")).Append('\t')
        .Append(GetSafeValue(lines, 3, "Plaka:")).Append('\t')
        .Append(GetSafeValue(lines, 4, "Seri:")).Append('\t');

    private static void AppendKurumselByIndex(StringBuilder sb, string[] lines) => sb
        .Append(GetSafeValue(lines, 0, "Sigortalı:")).Append("\t\t")
        .Append(GetSafeValue(lines, 1, "Vergi:")).Append("\t\t")
        .Append(GetSafeValue(lines, 2, "Plaka:")).Append('\t')
        .Append(GetSafeValue(lines, 3, "Seri:")).Append('\t');

    private static string GetSafeValue(string[] lines, int index, string prefix) =>
        index >= lines.Length ? "" : lines[index].Replace(prefix, "").Trim();

    private static string GetSafeValueByLabel(string[] lines, string prefix)
    {
        var line = lines.FirstOrDefault(l => l.Trim().StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        return line == null ? "" : line.Trim().Substring(prefix.Length).Trim();
    }
}