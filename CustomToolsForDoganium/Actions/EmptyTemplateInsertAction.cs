using System.Text;
using System.Windows.Forms;
using CustomToolsForDoganium.Helper;
using CustomToolsForDoganium.Hotkeys;

namespace CustomToolsForDoganium.Actions;

/// <summary>Ctrl+Ö(Oem2): boş bireysel müşteri bilgi şablonunu Notepad++'a yapıştırır.</summary>
internal sealed class EmptyTemplateInsertAction : IHotkeyAction
{
    public Keys VirtualKey => Keys.Oem2;
    public TargetApp SupportedApps => TargetApp.NotepadPlusPlus;

    public void Execute(string clipboardText, TargetApp activeApp, NotifyIcon notifyIcon)
    {
        var template = new StringBuilder()
            .AppendLine("###############################################################################")
            .AppendLine("Tel: ")
            .AppendLine()
            .AppendLine("Sigortalı: ")
            .AppendLine("TC: ")
            .AppendLine("DT: ")
            .AppendLine("Plaka: ")
            .AppendLine("Seri: ")
            .AppendLine("Motor: ")
            .AppendLine("Şasi: ")
            .AppendLine("Kullanım Şekli: ")
            .AppendLine("Model Yılı: ")
            .AppendLine("Marka Kodu: ")
            .AppendLine()
            .ToString();

        ClipboardPasteHelper.PasteAndRestore(template);
    }
}